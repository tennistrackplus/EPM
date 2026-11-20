/**
 * MetricViewImport
 * ------------------------------------------------------------------------
 * Reconstruye el modelo semántico de Draco ({ fact, fields }, el mismo
 * formato que LkmlImport) a partir del YAML de una Metric View de
 * Databricks. Es la función inversa de MetricViewExport.
 *
 * Lee la estructura con js-yaml (js/vendor/js-yaml.min.js) y los metadatos
 * de Draco de los comentarios de cada join / campo / medida («# draco-…»,
 * «# Jerarquía …», «# texto de la clave: …»). Con una Metric View escrita a
 * mano también funciona, en modo «best effort»:
 *   - source                 -> tabla de hechos (catálogo.esquema.tabla)
 *   - join con 'on' source.fk = join.clave (o using) -> dimensión «fk»
 *   - campo expr: join.col   -> atributo «col» de esa dimensión
 *   - campo expr: [source.]col -> dimensión de hechos «col»
 *   - medida AGG(col) [FILTER (WHERE col IN (...))] -> medida simple / filtrada
 *   - medida con MEASURE(a) +-* / MEASURE(b) -> medida calculada [a] - [b]
 * Lo que no encaja (joins anidados, expresiones SQL libres) se omite y se
 * avisa en «warnings».
 *
 * SEGURIDAD (como LkmlImport, auditoría F-07): catálogo, esquema, tabla y
 * columnas acaban dentro de las SQL del add-in: se exige que sean
 * identificadores simples. Del YAML NO se copia ninguna expresión SQL libre
 * al modelo: solo nombres de columnas, agregaciones de una lista cerrada y
 * valores de filtro (que el add-in ya escapa como literales).
 *
 * Función PURA (no toca Excel ni Office).
 */
(function () {
    "use strict";

    const IDENT_RE = /^[A-Za-z_][A-Za-z0-9_]{0,254}$/;
    const CATALOG_RE = /^[A-Za-z_][A-Za-z0-9_-]{0,254}$/;
    const AGGS = { SUM: "SUM", COUNT: "COUNT", AVG: "AVG", MIN: "MIN", MAX: "MAX" };

    function yamlLib() {
        const y = (typeof window !== "undefined" && window.jsyaml) || (typeof jsyaml !== "undefined" ? jsyaml : null);
        if (!y || typeof y.load !== "function") throw new Error("No se ha cargado el lector de YAML (js-yaml).");
        return y;
    }

    const unq = (s) => String(s == null ? "" : s).trim().replace(/^`(.*)`$/, "$1").replace(/``/g, "`");
    const fromB64 = (s) => JSON.parse(typeof atob === "function"
        ? decodeURIComponent(escape(atob(s)))
        : Buffer.from(s, "base64").toString("utf8"));

    /**
     * Comentarios de cada elemento de primer nivel de una sección
     * (joins / dimensions / fields / measures), en el mismo orden que la
     * lista que devuelve js-yaml: { joins: [[líneas], ...], ... }.
     */
    function collectComments(text) {
        const out = {};
        let section = null, itemIndent = -1, current = null;
        String(text).split(/\r?\n/).forEach(line => {
            const top = /^([A-Za-z_]+)\s*:/.exec(line);
            if (top) {
                section = top[1].toLowerCase();
                itemIndent = -1;
                current = null;
                if (!out[section]) out[section] = [];
                return;
            }
            if (!section) return;
            const dash = /^(\s*)-\s/.exec(line);
            if (dash) {
                const ind = dash[1].length;
                if (itemIndent === -1) itemIndent = ind;
                if (ind === itemIndent) {
                    current = [];
                    out[section].push(current);
                    return;
                }
            }
            const com = /^\s+#\s?(.*)$/.exec(line);
            if (com && current) current.push(com[1]);
        });
        return out;
    }

    function parseHierarchies(lines) {
        const text = lines.map(l => "# " + l).join("\n") + "\n";
        const res = [];
        const re = /#\s*Jerarquía\s*"([^"]+)"(?:\s*\(vac[ií]o:\s*"([^"]*)"\))?:\s*\n((?:[ \t]*#\s*\d+\.\s*.+\n?)+)/g;
        let m;
        while ((m = re.exec(text)) !== null) {
            const levels = [];
            const lre = /#\s*\d+\.\s*(.+)/g;
            let lm;
            while ((lm = lre.exec(m[3])) !== null) {
                const parts = lm[1].split(/\s*\|\s*texto:\s*/);
                const lvl = { attribute: parts[0].trim() };
                if (parts[1] && parts[1].trim()) lvl.textAttribute = parts[1].trim();
                levels.push(lvl);
            }
            if (levels.length) {
                const h = { name: m[1], levels };
                if (m[2] && m[2].trim()) h.emptyValue = m[2].trim();
                res.push(h);
            }
        }
        return res;
    }

    const meta = (lines, key) => {
        const re = new RegExp("^" + key + ":\\s*(.+?)\\s*$", "i");
        for (const l of lines || []) { const m = re.exec(l.trim()); if (m) return m[1]; }
        return null;
    };

    function splitTable(ref) {
        const parts = String(ref || "").trim().split(".").map(unq);
        if (parts.length === 3) return { project: parts[0], dataset: parts[1], table: parts[2] };
        if (parts.length === 2) return { project: "", dataset: parts[0], table: parts[1] };
        return { project: "", dataset: "", table: parts[0] || "" };
    }

    /** «source.fk = j.clave» (en cualquier orden) o «fk = clave» -> { fk, key } */
    function parseOn(on, joinName) {
        const s = String(on || "").trim();
        const m = /^\s*([`\w.]+)\s*=\s*([`\w.]+)\s*$/.exec(s);
        if (!m) return null;
        const side = (x) => {
            const p = x.split(".").map(unq);
            return p.length === 1 ? { t: "", c: p[0] } : { t: p[0], c: p[p.length - 1] };
        };
        const a = side(m[1]), b = side(m[2]);
        const jl = String(joinName).toLowerCase();
        if (a.t.toLowerCase() === "source" || b.t.toLowerCase() === jl) return { fk: a.c, key: b.c };
        if (b.t.toLowerCase() === "source" || a.t.toLowerCase() === jl) return { fk: b.c, key: a.c };
        return { fk: a.c, key: b.c };   // sin prefijos: izquierda = hechos, derecha = dimensión
    }

    /** AGG(col) / COUNT(DISTINCT col) [FILTER (WHERE col IN (...) [AND ...])] */
    function parseAggregate(expr) {
        const s = String(expr || "").trim();
        const m = /^(SUM|COUNT|AVG|MIN|MAX)\s*\(\s*(DISTINCT\s+)?([`\w]+)\s*\)\s*(?:FILTER\s*\(\s*WHERE\s+([\s\S]+)\)\s*)?$/i.exec(s);
        if (!m) return null;
        const agg = m[2] ? "COUNT_DISTINCT" : AGGS[m[1].toUpperCase()];
        if (m[2] && m[1].toUpperCase() !== "COUNT") return null;
        const out = { aggregation: agg, column: unq(m[3]), conditions: [] };
        if (m[4]) {
            for (const part of m[4].split(/\s+AND\s+/i)) {
                const c = /^\s*([`\w]+)\s+(NOT\s+)?IN\s*\(([\s\S]*)\)\s*$/i.exec(part);
                if (!c) return null;
                const vals = [];
                const vre = /'((?:[^'\\]|\\.|'')*)'|(-?\d+(?:\.\d+)?)/g;
                let vm;
                while ((vm = vre.exec(c[3])) !== null) {
                    vals.push(vm[1] !== undefined ? vm[1].replace(/''/g, "'").replace(/\\(.)/g, "$1") : vm[2]);
                }
                out.conditions.push({ column: unq(c[1]), dataType: /^-?\d/.test(vals[0] || "") && !/'/.test(c[3]) ? "NUMERIC" : "STRING",
                    op: c[2] ? "NOT_IN" : "IN", values: vals });
            }
        }
        return out;
    }

    /** MEASURE(a) - MEASURE(b) ... -> [a] - [b] (solo + - * / ( ) números, ABS, ROUND) */
    function measureFormula(expr) {
        let s = String(expr || "");
        if (!/MEASURE\s*\(/i.test(s)) return null;
        s = s.replace(/MEASURE\s*\(\s*([`\w]+)\s*\)/gi, (_, n) => "[" + unq(n) + "]")
            .replace(/NULLIF\s*\(\s*([\s\S]+?)\s*,\s*0\s*\)/gi, "$1")
            .replace(/try_divide\s*\(/gi, "DIVIDIR(");
        const bare = s.replace(/\[[^\]]+\]/g, "");
        if (!/^[\s\d.+\-*/(),;A-Za-z]*$/.test(bare)) return null;
        const words = bare.match(/[A-Za-z_]+/g) || [];
        if (words.some(w => !/^(ABS|ROUND|REDONDEAR|DIVIDIR)$/i.test(w))) return null;
        return s.replace(/,/g, ";").replace(/\s+/g, " ").trim();
    }

    function checkIdent(v, what, re) {
        const s = String(v == null ? "" : v);
        if (!(re || IDENT_RE).test(s)) throw new Error(`El fichero contiene un nombre de ${what} no válido: "${s.slice(0, 80)}".`);
    }

    /**
     * @param {string} content Texto del fichero .yaml / .yml
     * @returns {{ fact, fields, warnings: string[] }}
     */
    function parseContent(content) {
        if (!content || typeof content !== "string" || !content.trim()) throw new Error("El fichero está vacío.");
        let doc;
        try {
            doc = yamlLib().load(content);
        } catch (e) {
            throw new Error("El YAML no es válido: " + (e.reason || e.message));
        }
        if (!doc || typeof doc !== "object" || !doc.source) {
            throw new Error("No parece una Metric View de Databricks: falta «source».");
        }
        if (/^\s*select\b/i.test(String(doc.source))) {
            throw new Error("La Metric View usa una consulta SQL como «source»: Draco necesita una tabla (catálogo.esquema.tabla).");
        }
        const comments = collectComments(content);
        const warnings = [];
        const fact = splitTable(doc.source);

        /* ---- dimensiones con join ---- */
        const joins = Array.isArray(doc.joins) ? doc.joins : [];
        const dimByJoin = new Map();
        const dims = [];
        joins.forEach((j, i) => {
            const lines = (comments.joins || [])[i] || [];
            const name = String(j.name || "");
            const on = j.on !== undefined ? j.on : (j.true !== undefined ? j.true : null);   // 'on' sin comillas = true en YAML 1.1
            let rel = null;
            if (on) rel = parseOn(on, name);
            else if (j.using) {
                const u = Array.isArray(j.using) ? j.using[0] : String(j.using).split(",")[0];
                rel = { fk: unq(u), key: unq(u) };
            }
            if (!rel) { warnings.push(`Join «${name}»: no se entiende su condición de unión; se omite.`); return; }
            if (Array.isArray(j.joins) && j.joins.length) warnings.push(`Join «${name}»: los joins anidados (copo de nieve) no se importan.`);
            const t = splitTable(j.source);
            const d = {
                enabled: true, type: "DIMENSION",
                name: meta(lines, "draco-dimension") || rel.fk,
                relProject: t.project || fact.project, relDataset: t.dataset || fact.dataset, relTable: t.table,
                attributes: [], hierarchies: parseHierarchies(lines),
                textAttribute: meta(lines, "texto de la clave") || "",
                calcMembers: {},
                _key: rel.key
            };
            const cm = meta(lines, "draco-calculos");
            if (cm) { try { d.calcMembers = fromB64(cm); } catch (e) { /* comentario dañado */ } }
            dimByJoin.set(name.toLowerCase(), d);
            dims.push(d);
        });

        /* ---- campos -> atributos / dimensiones de hechos ---- */
        const fieldList = Array.isArray(doc.fields) ? doc.fields : (Array.isArray(doc.dimensions) ? doc.dimensions : []);
        const fieldComments = comments.fields || comments.dimensions || [];
        fieldList.forEach((f, i) => {
            const lines = fieldComments[i] || [];
            const expr = String(f.expr == null ? "" : f.expr).trim();
            const calc = meta(lines, "draco-atributo");
            let joinName = null, col = null;
            const m = /^([`\w]+)(?:\.([`\w]+))?$/.exec(expr);
            if (m && m[2]) { joinName = unq(m[1]); col = unq(m[2]); }
            else if (m) { col = unq(m[1]); }
            let attrName = col, expression = null;
            if (calc) {
                try { const d = fromB64(calc); attrName = d.name; expression = String(d.expression || ""); } catch (e) { /* dañado */ }
            }
            if (!attrName) { warnings.push(`Campo «${f.name}»: su expresión no es una columna; se omite.`); return; }
            const attr = { name: attrName, alias: f.display_name ? String(f.display_name) : attrName,
                dataType: meta(lines, "draco-tipo") || "", isKey: false, enabled: true };
            if (expression !== null) { attr.calc = true; attr.expression = expression; }

            const target = joinName && joinName.toLowerCase() !== "source" ? dimByJoin.get(joinName.toLowerCase()) : null;
            if (joinName && joinName.toLowerCase() !== "source" && !target) {
                warnings.push(`Campo «${f.name}»: usa el join «${joinName}», que no se ha importado; se omite.`);
                return;
            }
            if (target) {
                if (!target.attributes.some(a => a.name === attr.name)) target.attributes.push(attr);
                if (target._order === undefined) target._order = i;
                return;
            }
            // Dimensión de la tabla de hechos: una por columna (o la que diga draco-dimension)
            const dname = meta(lines, "draco-dimension") || attrName;
            let d = dims.find(x => !x.relTable && x.name === dname);
            if (!d) {
                d = { enabled: true, type: "DIMENSION", name: dname, relProject: "", relDataset: "", relTable: "",
                    attributes: [], hierarchies: parseHierarchies(lines), textAttribute: meta(lines, "texto de la clave") || "",
                    calcMembers: {}, _key: dname, _order: i };
                const cm = meta(lines, "draco-calculos");
                if (cm) { try { d.calcMembers = fromB64(cm); } catch (e) { /* dañado */ } }
                dims.push(d);
            }
            d.attributes.push(attr);
        });
        // Mismo orden que en Draco: el de la lista de campos (los joins sin campos, al final)
        dims.sort((a, b) => (a._order === undefined ? 1e9 : a._order) - (b._order === undefined ? 1e9 : b._order));
        dims.forEach(d => {
            delete d._order;
            let key = d.attributes.find(a => a.name === d._key && !a.calc);
            if (!key && d.relTable) {
                key = { name: d._key, alias: d._key, dataType: "", isKey: false, enabled: true };
                d.attributes.unshift(key);
            }
            if (key) key.isKey = true;
            delete d._key;
        });

        /* ---- medidas ---- */
        const measures = [];
        const mlist = Array.isArray(doc.measures) ? doc.measures : [];
        mlist.forEach((m, i) => {
            const lines = (comments.measures || [])[i] || [];
            const name = String(m.name || "");
            const def = meta(lines, "draco-medida");
            if (def) {
                try {
                    const d = fromB64(def);
                    if (d && (d.measureKind === "CALC" || d.measureKind === "FILTERED") && IDENT_RE.test(d.name || "")) {
                        measures.push({ enabled: true, type: "MEASURE", name: d.name, alias: d.alias || d.name, dataType: "NUMERIC",
                            measureKind: d.measureKind, formula: String(d.formula || ""), base: String(d.base || ""),
                            conditions: Array.isArray(d.conditions) ? d.conditions : [],
                            aggregation: d.aggregation || "SUM", format: d.format || "Auto" });
                        return;
                    }
                } catch (e) { /* comentario dañado: se interpreta la expresión */ }
            }
            const alias = m.display_name ? String(m.display_name) : name;
            const fmt = meta(lines, "draco-formato") || "Auto";
            const ag = parseAggregate(m.expr);
            if (ag && !ag.conditions.length) {
                if (ag.column.toLowerCase() !== name.toLowerCase()) {
                    // Draco identifica la medida simple por su columna: si la Metric View
                    // la llama distinto, se conserva su nombre como alias.
                    measures.push({ enabled: true, type: "MEASURE", name: ag.column, alias: name, aggregation: ag.aggregation, format: fmt });
                } else {
                    measures.push({ enabled: true, type: "MEASURE", name: ag.column, alias: alias !== name ? alias : undefined, aggregation: ag.aggregation, format: fmt });
                }
                return;
            }
            if (ag && ag.conditions.length) {
                if (!measures.some(x => !x.measureKind && x.name.toLowerCase() === ag.column.toLowerCase())) {
                    measures.push({ enabled: true, type: "MEASURE", name: ag.column, aggregation: "SUM", format: "Auto" });
                }
                measures.push({ enabled: true, type: "MEASURE", name, alias, dataType: "NUMERIC", measureKind: "FILTERED",
                    formula: "", base: ag.column, conditions: ag.conditions, aggregation: ag.aggregation, format: fmt });
                return;
            }
            const formula = measureFormula(m.expr);
            if (formula) {
                measures.push({ enabled: true, type: "MEASURE", name, alias, dataType: "NUMERIC", measureKind: "CALC",
                    formula, base: "", conditions: [], aggregation: "SUM", format: fmt });
                return;
            }
            warnings.push(`Medida «${name}»: expresión no soportada por Draco (${String(m.expr).slice(0, 60)}); se omite.`);
        });
        measures.forEach(x => { if (x.alias === undefined) delete x.alias; });

        if (!fact.table && !dims.length && !measures.length) {
            throw new Error("No se ha podido reconocer ninguna tabla de hechos, dimensión o medida en el fichero.");
        }

        /* ---- validación de identificadores (F-07) ---- */
        if (fact.project) checkIdent(fact.project, "catálogo", CATALOG_RE);
        if (fact.dataset) checkIdent(fact.dataset, "esquema");
        checkIdent(fact.table, "tabla");
        dims.forEach(d => {
            checkIdent(d.name, "campo");
            if (d.relProject) checkIdent(d.relProject, "catálogo", CATALOG_RE);
            if (d.relDataset) checkIdent(d.relDataset, "esquema");
            if (d.relTable) checkIdent(d.relTable, "tabla");
            d.attributes.forEach(a => checkIdent(a.name, "atributo"));
            if (d.textAttribute) checkIdent(d.textAttribute, "atributo de texto");
            d.hierarchies.forEach(h => h.levels.forEach(l => {
                checkIdent(l.attribute, "nivel de jerarquía");
                if (l.textAttribute) checkIdent(l.textAttribute, "atributo de texto");
            }));
        });
        measures.forEach(x => {
            checkIdent(x.name, "medida");
            if (x.base) checkIdent(x.base, "medida base");
            (x.conditions || []).forEach(c => checkIdent(c.column, "columna de filtro"));
        });

        const result = { fact, fields: [...dims, ...measures], warnings };
        if (/^[ \t]*#[ \t]*draco[ \t]*:[ \t]*planning[ \t]*$/im.test(content)) result.dracoPlanning = true;
        return result;
    }

    window.MetricViewImport = { parseContent };
})();
