/**
 * ============================================================
 * EPM ADD-IN — CONSULTAS DAX CONTRA MODELOS SEMÁNTICOS DE POWER BI
 * ============================================================
 * Los modelos importados desde Power BI/Fabric (source.type =
 * "fabric-semantic-model", ver fabricSemanticModel.js) NO se consultan con
 * SQL: las relaciones, las medidas DAX y la seguridad por filas las
 * resuelve el propio modelo. Este módulo sustituye, SOLO para esos modelos,
 * a buildSQL() + executeSQL() del informe Dinámico:
 *
 *   1. buildDynamicQuery(): a partir de ReportState (Filas, Columnas,
 *      Medidas, Filtros) genera UNA consulta DAX. Cada "grouping set" que
 *      generaría buildSQL (jerarquías y subtotales) se convierte en un
 *      SUMMARIZECOLUMNS, y todos se unen con UNION. Así las medidas DAX
 *      (no aditivas) las calcula siempre Power BI en cada nivel.
 *   2. execute(): la lanza contra la API executeQueries de Power BI y
 *      reproduce en JS la parte final del SQL (filtro de filas "todo NULL"
 *      por dimensión, etiquetas 'TOTAL', ROW_ID/COLUMN_ID con DENSE_RANK y
 *      ORDER BY), devolviendo EL MISMO texto "pseudo JSON de BigQuery" que
 *      ya leen jsonTo3MatricesCore/parseJsonValueTriples. El pintado en
 *      Excel no cambia en absoluto.
 *
 * Columnas del resultado (igual que el SELECT final de buildSQL):
 *   ROW_ID, COLUMN_ID, <atributos de Columnas>, <atributos de Filas>, <medidas>
 * (ROW_ID se calcula con el eje Columnas y COLUMN_ID con el eje Filas,
 * igual que el VBA original.)
 */
(function () {

    const PBI_API = "https://api.powerbi.com/v1.0/myorg";

    // fabric.js declara "const FB" en el ámbito global del script: es
    // accesible por su nombre, pero NO como window.FB.
    function fabricSession() {
        if (typeof FB !== "undefined") return FB;
        return window.FB || null;
    }

    // -----------------------------------------------------------------
    // Modelo activo
    // -----------------------------------------------------------------
    function activeModel() {
        const store = window.SemanticModelStore;
        if (!store) return null;
        const name = store.getActiveModelName();
        return name ? store.getModel(name) : null;
    }

    /** true si el modelo semántico activo es un modelo de Power BI/Fabric */
    function isActiveModel() {
        const m = activeModel();
        return !!(m && m.source && m.source.type === "fabric-semantic-model");
    }

    // -----------------------------------------------------------------
    // Lectura de los grids del modelo (fila 0 = cabecera)
    // -----------------------------------------------------------------
    const up = v => String(v === null || v === undefined ? "" : v).trim().toUpperCase();

    /** MODEL_ATRIBUTES: DIMENSION(1) ATRIBUTE(2) DIM_TABLE(5) DIM_FIELD(6) DISPLAY_NAME(7) DATA_TYPE(8) */
    function findAttribute(atributesGrid, dimension, attributeName) {
        const rows = (atributesGrid && atributesGrid.values) || [];
        const d = up(dimension), a = up(attributeName);
        let byAlias = null;
        for (let r = 1; r < rows.length; r++) {
            const row = rows[r];
            if (up(row[1]) !== d) continue;
            const info = { table: String(row[5]), column: String(row[6]), dataType: up(row[8]) };
            if (up(row[2]) === a) return info;
            if (!byAlias && up(row[7]) === a) byAlias = info;
        }
        if (byAlias) return byAlias;
        throw new Error(`El atributo "${attributeName}" de "${dimension}" no existe en el modelo semántico activo.`);
    }

    /** MODEL_MEASURES: MEASURE(1) FACT_TABLE(4) FACT_FIELD(5) AGGREGATION(6) */
    function findMeasure(measuresGrid, measureName) {
        const rows = (measuresGrid && measuresGrid.values) || [];
        const n = up(measureName);
        for (let r = 1; r < rows.length; r++) {
            const row = rows[r];
            if (up(row[1]) === n) {
                return { table: String(row[4]), field: String(row[5]), aggregation: up(row[6]) };
            }
        }
        throw new Error(`La medida "${measureName}" no existe en el modelo semántico activo.`);
    }

    // -----------------------------------------------------------------
    // Sintaxis DAX
    // -----------------------------------------------------------------
    const daxTable = t => "'" + String(t).replace(/'/g, "''") + "'";
    const daxColumn = (t, c) => daxTable(t) + "[" + String(c).replace(/]/g, "]]") + "]";
    const daxMeasureRef = m => "[" + String(m).replace(/]/g, "]]") + "]";
    const daxString = s => '"' + String(s).replace(/"/g, '""') + '"';

    const NUMERIC_TYPES = ["INTEGER", "INT64", "FLOAT", "NUMERIC", "BIGNUMERIC", "DECIMAL", "DOUBLE"];

    function daxLiteral(value, dataType) {
        const s = String(value).trim();
        if (NUMERIC_TYPES.includes(dataType) && s !== "" && !isNaN(Number(s))) return String(Number(s));
        if (dataType === "BOOLEAN") return /^(true|1|verdadero|x)$/i.test(s) ? "TRUE()" : "FALSE()";
        if (dataType === "DATETIME" || dataType === "DATE") {
            const m = s.match(/^(\d{4})-(\d{2})-(\d{2})(?:[T ](\d{2}):(\d{2})(?::(\d{2}))?)?/);
            if (m) {
                const date = `DATE(${+m[1]}, ${+m[2]}, ${+m[3]})`;
                return m[4] ? `(${date} + TIME(${+m[4]}, ${+m[5]}, ${+(m[6] || 0)}))` : date;
            }
        }
        return daxString(s);
    }

    function measureExpression(measuresGrid, measureName) {
        const m = findMeasure(measuresGrid, measureName);
        const col = () => daxColumn(m.table, m.field);
        switch (m.aggregation) {
            case "DAX": return daxMeasureRef(m.field || measureName);
            case "SUM": return `SUM(${col()})`;
            case "AVG":
            case "AVERAGE": return `AVERAGE(${col()})`;
            case "MIN": return `MIN(${col()})`;
            case "MAX": return `MAX(${col()})`;
            case "COUNT": return `COUNTA(${col()})`;
            case "COUNT_DISTINCT":
            case "COUNTDISTINCT":
            case "DISTINCTCOUNT": return `DISTINCTCOUNT(${col()})`;
            default:
                throw new Error(`Agregación "${m.aggregation}" no soportada en la medida "${measureName}".`);
        }
    }

    // -----------------------------------------------------------------
    // Filtros -> tablas de filtro de SUMMARIZECOLUMNS
    // (misma semántica que buildFilterConditions en commands.js)
    // -----------------------------------------------------------------
    function parseStoredFilter(raw) {
        const s = String(raw || "").trim();
        if (s === "") return null;
        if (s[0] === "{") {
            try {
                const parsed = JSON.parse(s);
                if (parsed && typeof parsed === "object" && parsed.mode) return parsed;
            } catch (e) { /* valor simple */ }
        }
        return { mode: "values", include: true, values: [s] };
    }

    const itemValue = it => (it && typeof it === "object") ? (it.value !== undefined ? it.value : it.display) : it;

    /**
     * Construye el predicado DAX de un filtro. Devuelve { predicate, columns }
     * (columns = columnas DAX que usa, todas de la misma tabla) o null.
     */
    function buildFilterPredicate(atributesGrid, filterRow) {
        const filter = parseStoredFilter(filterRow.Value);
        if (!filter) return null;

        const columns = new Map(); // "T|C" -> daxColumn
        const tables = new Set();
        const colFor = attr => {
            const info = findAttribute(atributesGrid, filterRow.Dimension, attr);
            tables.add(info.table);
            const ref = daxColumn(info.table, info.column);
            columns.set(info.table + "|" + info.column, ref);
            return { ref, type: info.dataType };
        };

        const valuesCond = (attr, values, include) => {
            const vals = (values || []).map(v => String(v)).filter(v => v.trim() !== "");
            if (vals.length === 0) return "";
            const { ref, type } = colFor(attr);
            const list = vals.map(v => daxLiteral(v, type));
            const cond = vals.length === 1 ? `${ref} = ${list[0]}` : `${ref} IN {${list.join(", ")}}`;
            return include ? cond : `NOT (${cond})`;
        };

        const rangeCond = (attr, from, to, include) => {
            const { ref, type } = colFor(attr);
            const cond = `(${ref} >= ${daxLiteral(from, type)} && ${ref} <= ${daxLiteral(to, type)})`;
            return include ? cond : `NOT ${cond}`;
        };

        const or = parts => parts.length === 1 ? parts[0] : "(" + parts.join(" || ") + ")";
        const and = parts => parts.length === 1 ? parts[0] : "(" + parts.join(" && ") + ")";

        const groupedItems = (items, include) => {
            if (!items || items.length === 0) return "";
            const byAttr = new Map();
            items.forEach(it => {
                const attr = it && it.attribute;
                if (!attr) return;
                if (!byAttr.has(attr)) byAttr.set(attr, []);
                byAttr.get(attr).push(itemValue(it));
            });
            const parts = [];
            byAttr.forEach((vals, attr) => { const c = valuesCond(attr, vals, true); if (c) parts.push(c); });
            if (parts.length === 0) return "";
            return include ? or(parts) : `NOT ${or(parts)}`;
        };

        let predicate = "";
        const isHierarchy = (filter.items && filter.items.length) || (filter.excludeItems && filter.excludeItems.length);

        if (isHierarchy) {
            if (filter.mode === "list") {
                predicate = and([groupedItems(filter.items, true), groupedItems(filter.excludeItems, false)].filter(Boolean));
            } else {
                const inc = groupedItems(filter.items, true);
                predicate = inc ? (filter.include === false ? `NOT ${inc}` : inc) : "";
            }
        } else {
            const attr = filterRow.AttributeName;
            if (filter.mode === "list") {
                const included = [];
                const incValues = valuesCond(attr, filter.values, true);
                if (incValues) included.push(incValues);
                (filter.ranges || []).forEach(r => included.push(rangeCond(attr, r.from, r.to, true)));
                const parts = [];
                if (included.length) parts.push(or(included));
                const excValues = valuesCond(attr, filter.excludeValues, false);
                if (excValues) parts.push(excValues);
                (filter.excludeRanges || []).forEach(r => parts.push(rangeCond(attr, r.from, r.to, false)));
                predicate = parts.length ? and(parts) : "";
            } else if (filter.mode === "range") {
                predicate = rangeCond(attr, filter.from, filter.to, filter.include !== false);
            } else if (filter.mode === "mixed") {
                const parts = [
                    valuesCond(attr, filter.values, filter.valuesInclude !== false),
                    (filter.from !== undefined && filter.to !== undefined) ? rangeCond(attr, filter.from, filter.to, filter.rangeInclude !== false) : ""
                ].filter(Boolean);
                predicate = parts.length ? or(parts) : "";
            } else {
                predicate = valuesCond(attr, filter.values, filter.include !== false);
            }
        }

        if (!predicate) return null;
        if (tables.size > 1) {
            throw new Error(`El filtro de "${filterRow.Dimension}" mezcla columnas de varias tablas, algo que no se puede aplicar a un modelo de Power BI.`);
        }
        return { predicate, columns: Array.from(columns.values()) };
    }

    // -----------------------------------------------------------------
    // Grouping sets (misma lógica que computeGroupingSetDimensionsAndHierarchies
    // + buildConfigSets/buildConfigSetsWithSubtotals, pero con índices)
    // -----------------------------------------------------------------
    function hierarchyRunLengths(fields) {
        const out = [];
        let lastDim = "", lastLevel = 0, count = 0;
        fields.forEach(f => {
            const level = Number(f.Hierarchy) || 0;
            if (count > 0 && f.Dimension === lastDim && level === lastLevel + 1) {
                count++;
            } else {
                if (count > 0) out.push(count);
                count = 1;
            }
            lastDim = f.Dimension;
            lastLevel = level;
        });
        if (count > 0) out.push(count);
        return out;
    }

    function hierarchyCombinations(hierarchies) {
        if (hierarchies.length === 0) return [[]];
        const result = [];
        const current = hierarchies.slice();
        while (true) {
            result.push(current.slice());
            let i = current.length - 1;
            let done = false;
            while (true) {
                current[i] -= 1;
                if (current[i] > 0) break;
                current[i] = hierarchies[i];
                i -= 1;
                if (i < 0) { done = true; break; }
            }
            if (done) break;
        }
        return result;
    }

    /**
     * Devuelve la lista de grouping sets como arrays de claves de campo:
     * "R<i>" (i-ésimo campo de Filas) y "C<k>" (k-ésimo de Columnas).
     */
    function groupingSets(state, withSubtotals) {
        const dims = [
            ...state.Rows.map((f, i) => "R" + i),
            ...state.Columns.map((f, k) => "C" + k)
        ];
        const hierarchies = [...hierarchyRunLengths(state.Rows), ...hierarchyRunLengths(state.Columns)];

        const sets = [];
        hierarchyCombinations(hierarchies).forEach(item => {
            const keys = [];
            let offset = 0;
            item.forEach((len, i) => {
                for (let j = 0; j < len; j++) keys.push(dims[offset + j]);
                offset += hierarchies[i];
            });
            sets.push(keys);
        });

        if (withSubtotals) extraSubtotalSets(state).forEach(k => sets.push(k));
        return sets;
    }

    /** Conjuntos extra de los subtotales marcados en L/R (buildConfigSetsWithSubtotals) */
    function extraSubtotalSets(state) {
        const sets = [];
        const colKeys = state.Columns.map((f, k) => "C" + k);
        const rowKeys = state.Rows.map((f, i) => "R" + i);
        state.Columns.forEach((c, idx) => {
            if (c.Subtotal) sets.push([...colKeys.slice(0, idx), ...rowKeys]);
        });
        state.Rows.forEach((r, idx) => {
            if (r.Subtotal) sets.push([...colKeys, ...rowKeys.slice(0, idx)]);
        });
        return sets;
    }

    /**
     * Tramos de jerarquía de un eje como listas de claves ("R0","R1"...):
     * mismo criterio que hierarchyRunLengths (misma dimensión y NIVEL
     * consecutivo).
     */
    function hierarchyRuns(fields, prefix) {
        const lengths = hierarchyRunLengths(fields);
        const runs = [];
        let i = 0;
        lengths.forEach(len => {
            runs.push(fields.slice(i, i + len).map((f, j) => prefix + (i + j)));
            i += len;
        });
        return runs;
    }

    // -----------------------------------------------------------------
    // 1) Construcción de la consulta
    // -----------------------------------------------------------------
    /**
     * @param {object} state         ReportState (tras loadReportDefinition + applyFilterMeasuresToDynamicReport)
     * @param {object} measuresGrid  SemanticModelStore.getModelGrid("MODEL_MEASURES")
     * @param {object} atributesGrid SemanticModelStore.getModelGrid("MODEL_ATRIBUTES")
     * @param {boolean} subtotalsOnTop EDIT_REPORT!D4
     * @returns {{ kind:"powerbi-dax", text:string, plan:object }}
     */
    function buildDynamicQuery(state, measuresGrid, atributesGrid, subtotalsOnTop) {
        const model = activeModel();
        if (!model || !model.source || model.source.type !== "fabric-semantic-model") {
            throw new Error("El modelo semántico activo no es un modelo de Power BI.");
        }
        if (!state.Measures || state.Measures.length === 0) {
            throw new Error("Añade al menos una medida al informe.");
        }

        const withSubtotals = state.Columns.some(c => c.Subtotal) || state.Rows.some(r => r.Subtotal);

        // Campos en el orden del SELECT final: Columnas y luego Filas
        const fields = [
            ...state.Columns.map((f, k) => ({ key: "C" + k, axis: "cols", def: f })),
            ...state.Rows.map((f, i) => ({ key: "R" + i, axis: "rows", def: f }))
        ];
        fields.forEach((f, idx) => {
            const info = findAttribute(atributesGrid, f.def.Dimension, f.def.AttributeName);
            f.ref = daxColumn(info.table, info.column);
            f.dataType = info.dataType;
            f.alias = "c" + idx;
        });
        const byKey = new Map(fields.map(f => [f.key, f]));

        const measures = state.Measures.map((m, idx) => ({
            name: m.Name,
            alias: "m" + idx,
            expr: measureExpression(measuresGrid, m.Name)
        }));

        // Filtros del informe (zona Filtros + filtros "bloqueados")
        const filters = (state.Filters || [])
            .filter(f => String(f.Value).trim() !== "")
            .map(f => buildFilterPredicate(atributesGrid, f))
            .filter(Boolean);

        let dax = "";
        if (filters.length > 0) {
            dax += "DEFINE\n";
            filters.forEach((f, i) => {
                dax += `    VAR __F${i} = FILTER(ALL(${f.columns.join(", ")}), ${f.predicate})\n`;
            });
        }

        const filterArgs = filters.map((f, i) => `__F${i}`);
        const measureArgs = measures.map(m => `${daxString(m.alias)}, ${m.expr}`);
        const measureCols = measures.map(m => `${daxString(m.alias)}, ${daxMeasureRef(m.alias)}`);

        // Cada parte del UNION devuelve las mismas columnas: un indicador
        // "fN" por campo (1 = ese campo está totalizado en esta fila, como
        // GROUPING() en SQL), su valor "cN" y las medidas "mN".
        const wrap = (args, cols) =>
            `    SELECTCOLUMNS(\n        SUMMARIZECOLUMNS(\n            ${args.join(",\n            ")}\n        ),\n        ${cols.join(",\n        ")}\n    )`;

        // Grouping set "plano" (sin subtotales internos): los campos que no
        // están en keys van a BLANK() con su indicador a 1.
        const plainPart = keys => {
            const groupBy = [];
            keys.forEach(k => { const ref = byKey.get(k).ref; if (!groupBy.includes(ref)) groupBy.push(ref); });
            const inSet = new Set(keys);
            const cols = [
                ...fields.map(f => `"f${f.alias.slice(1)}", ${inSet.has(f.key) ? 0 : 1}`),
                ...fields.map(f => `${daxString(f.alias)}, ${inSet.has(f.key) ? f.ref : "BLANK()"}`),
                ...measureCols
            ];
            return wrap([...groupBy, ...filterArgs, ...measureArgs], cols);
        };

        const refs = fields.map(f => f.ref);
        const hasDuplicateRefs = new Set(refs).size !== refs.length;
        const parts = [];

        if (!hasDuplicateRefs) {
            // Jerarquías con ROLLUPADDISSUBTOTAL: UNA sola agrupación calcula
            // todos los niveles de todas las jerarquías y sus combinaciones
            // (los mismos grouping sets que genera buildSQL). En cada tramo de
            // jerarquía el primer nivel va como agrupación normal (Draco nunca
            // pide el total general de una jerarquía) y el resto dentro de
            // ROLLUPADDISSUBTOTAL, que los va totalizando de abajo arriba.
            const runs = [...hierarchyRuns(state.Rows, "R"), ...hierarchyRuns(state.Columns, "C")];
            const groupArgs = [];
            const flagExpr = new Map(); // key -> expresión del indicador
            runs.forEach((keys, r) => {
                groupArgs.push(byKey.get(keys[0]).ref);
                flagExpr.set(keys[0], "0");
                if (keys.length > 1) {
                    const pairs = keys.slice(1).map((k, i) => {
                        const flagName = `__S${r}_${i}`;
                        flagExpr.set(k, `IF([${flagName}], 1, 0)`);
                        return `${byKey.get(k).ref}, ${daxString(flagName)}`;
                    });
                    groupArgs.push(`ROLLUPADDISSUBTOTAL(${pairs.join(", ")})`);
                }
            });
            const cols = [
                ...fields.map(f => `"f${f.alias.slice(1)}", ${flagExpr.get(f.key)}`),
                ...fields.map(f => `${daxString(f.alias)}, ${f.ref}`),
                ...measureCols
            ];
            parts.push(wrap([...groupArgs, ...filterArgs, ...measureArgs], cols));

            // Subtotales marcados en L/R (prefijos de un eje x el otro eje
            // completo): conjuntos adicionales, igual que buildConfigSetsWithSubtotals.
            if (withSubtotals) extraSubtotalSets(state).forEach(keys => parts.push(plainPart(keys)));
        } else {
            // Un mismo campo repetido en el informe: no se puede agrupar dos
            // veces por la misma columna en un SUMMARIZECOLUMNS, así que se
            // vuelve al método de un grouping set por parte.
            groupingSets(state, withSubtotals).forEach(keys => parts.push(plainPart(keys)));
        }

        dax += "EVALUATE\n" + (parts.length === 1 ? parts[0].replace(/^ {4}/gm, "") : "UNION(\n" + parts.join(",\n") + "\n)");

        return {
            kind: "powerbi-dax",
            mode: "dynamic",
            text: dax,
            plan: {
                workspaceId: model.source.workspaceId,
                semanticModelId: model.source.semanticModelId,
                withSubtotals,
                subtotalsOnTop: !!subtotalsOnTop,
                fields: fields.map(f => ({
                    key: f.key, axis: f.axis, alias: f.alias, dataType: f.dataType,
                    dimension: f.def.Dimension, hierarchy: Number(f.def.Hierarchy) || 0, subtotal: !!f.def.Subtotal
                })),
                measures: measures.map(m => ({ alias: m.alias, name: m.name }))
            }
        };
    }

    // -----------------------------------------------------------------
    // 1b) Valores de un atributo o jerarquía (diálogos de filtro y
    //     selector de miembros). Sustituye a buildAttributeSQL /
    //     buildHierarchySQL de excelService.js.
    // -----------------------------------------------------------------
    /** MODEL_HIER: HIERARCHY(1) NIVEL(2) DIMENSION(3) DIM_TABLE(7) DIM_FIELD(8) */
    function findHierarchyLevels(hierGrid, dimension, hierarchy) {
        const rows = (hierGrid && hierGrid.values) || [];
        const levels = [];
        for (let r = 1; r < rows.length; r++) {
            const row = rows[r];
            if (up(row[3]) === up(dimension) && up(row[1]) === up(hierarchy)) {
                levels.push({ level: Number(row[2]) || 0, table: String(row[7]), column: String(row[8]) });
            }
        }
        return levels.sort((a, b) => a.level - b.level);
    }

    /**
     * Consulta de los valores de "name" (atributo o jerarquía) de una
     * dimensión, o "" si no existe (mismo contrato que buildFilterValuesSQL).
     */
    function buildMembersQuery(dimension, name, atributesGrid, hierGrid) {
        const model = activeModel();
        const source = model && model.source;

        let columns = null;
        try {
            const info = findAttribute(atributesGrid, dimension, name);
            columns = [{ table: info.table, column: info.column }];
        } catch (e) {
            const levels = findHierarchyLevels(hierGrid, dimension, name);
            if (levels.length > 0) columns = levels;
        }
        if (!columns) return "";

        const refs = columns.map(c => daxColumn(c.table, c.column));
        const text = "EVALUATE\n" +
            (refs.length === 1 ? `DISTINCT(${refs[0]})` : `SUMMARIZECOLUMNS(${refs.join(", ")})`) +
            "\nORDER BY " + refs.join(", ");

        return {
            kind: "powerbi-dax",
            mode: "members",
            text,
            plan: {
                workspaceId: source.workspaceId,
                semanticModelId: source.semanticModelId,
                names: columns.map(c => c.column)
            }
        };
    }

    /**
     * Mismo texto que un resultado de BigQuery con su "schema": loadJsonTree
     * (filterModal.js) y parseMemberJsonTree (commands.js) leen los nombres
     * de los campos ("name") y después los valores ("v") entre comillas.
     * Los valores en blanco se envían como "" para no desalinear las columnas.
     */
    function membersToPseudoBqJson(rows, names) {
        const fields = names.map(n => '{"name": "' + String(n).replace(/"/g, "") + '"}').join(",");
        const out = rows.map(row => {
            const keys = Object.keys(row);
            const vals = names.map((n, i) => {
                const v = keys[i] !== undefined ? row[keys[i]] : null;
                const text = isNull(v) ? "" : String(v).replace(/\\/g, "\\\\").replace(/"/g, "'");
                return '{"v": "' + text + '"}';
            });
            return '{"f":[' + vals.join(",") + "]}";
        });
        return '{"schema":{"fields":[' + fields + ']},"rows":[' + out.join(",") + "]}";
    }

    // -----------------------------------------------------------------
    // 1c) Modo Fijo (Filas y Columnas estáticas: cabeceras EPM_VALUE).
    //     Sustituye a buildSQLFixed: una celda por ROW_ID x COLUMN_ID.
    // -----------------------------------------------------------------
    const DATATABLE_TYPES = {
        INTEGER: "INTEGER", INT64: "INTEGER",
        FLOAT: "DOUBLE", DOUBLE: "DOUBLE", NUMERIC: "DOUBLE", BIGNUMERIC: "DOUBLE", DECIMAL: "DOUBLE",
        BOOLEAN: "BOOLEAN", DATETIME: "DATETIME", DATE: "DATETIME"
    };

    function datatableLiteral(value, dtType) {
        const s = String(value === null || value === undefined ? "" : value).trim();
        if ((dtType === "INTEGER" || dtType === "DOUBLE") && s !== "" && !isNaN(Number(s))) return String(Number(s));
        if (dtType === "BOOLEAN") return /^(true|1|verdadero|x)$/i.test(s) ? "TRUE" : "FALSE";
        return daxString(s);
    }

    /**
     * Agrupa las definiciones de un eje (una por celda EPM_VALUE) por su
     * ROW_ID/COLUMN_ID y, después, por la combinación de columnas que
     * filtran ("firma"; normalmente todas las del eje). Cada firma es un
     * bloque: una DATATABLE con una fila por ID y sus valores.
     */
    function axisBlocks(defs, atributesGrid, idName, prefix) {
        const byId = new Map();
        defs.forEach(d => {
            const id = Number(d.R);
            const info = findAttribute(atributesGrid, d.Dimension, d.AttributeName);
            if (!byId.has(id)) byId.set(id, new Map());
            byId.get(id).set(info.table + "|" + info.column, { info, value: d.Value });
        });

        const bySignature = new Map();
        byId.forEach((conds, id) => {
            const sig = Array.from(conds.keys()).sort().join("||");
            if (!bySignature.has(sig)) bySignature.set(sig, []);
            bySignature.get(sig).push({ id, conds });
        });

        const blocks = [];
        bySignature.forEach((items, sig) => {
            const b = blocks.length;
            const keys = sig.split("||");
            const cols = keys.map((k, i) => {
                const { info } = items[0].conds.get(k);
                return {
                    key: k,
                    name: prefix + b + "_" + i,
                    ref: daxColumn(info.table, info.column),
                    dtType: DATATABLE_TYPES[info.dataType] || "STRING"
                };
            });
            const header = [daxString(idName), "INTEGER", ...cols.flatMap(c => [daxString(c.name), c.dtType])].join(", ");
            const dataRows = items.map(it =>
                "{" + [String(it.id), ...keys.map((k, i) => datatableLiteral(it.conds.get(k).value, cols[i].dtType))].join(", ") + "}");
            blocks.push({
                varName: "__" + prefix.toUpperCase() + b,
                table: `DATATABLE(${header}, {${dataRows.join(", ")}})`,
                cols,
                items: items.map(it => ({ id: it.id, values: keys.map((k, i) => normKey(it.conds.get(k).value, cols[i].dtType)) }))
            });
        });
        return blocks;
    }

    /** Normaliza un valor para comparar cabeceras (EPM_VALUE) con resultados del DAX */
    function normKey(v, dtType) {
        if (isNull(v)) return "";
        if (dtType === "INTEGER" || dtType === "DOUBLE") {
            const n = Number(v);
            return isNaN(n) ? String(v).trim() : String(n);
        }
        if (dtType === "BOOLEAN") return /^(true|1|verdadero|x)$/i.test(String(v)) ? "true" : "false";
        if (dtType === "DATETIME") return String(v).trim().replace(/T00:00:00(\.0+)?Z?$/, "").replace(" 00:00:00", "");
        return String(v);
    }

    /**
     * Modo Fijo agregado, como el SQL (un escaneo por grupo, nada de celda
     * a celda): por cada combinación de bloque de filas x bloque de
     * columnas, UN SUMMARIZECOLUMNS agrupado por las columnas de ambos
     * bloques y filtrado (TREATAS) a las combinaciones de valores de sus
     * cabeceras. Como cada fila del resultado corresponde exactamente a una
     * combinación ROW_ID x COLUMN_ID, el valor es exacto para cualquier
     * medida (también DAX, promedios o recuentos distintos). El reparto a
     * ROW_ID/COLUMN_ID se hace después en JS (fixedToPseudoBqJson).
     *
     * @param rowsDefs/colDefs definiciones leídas de las fórmulas EPM_VALUE
     *        (readRowDefinitions/readColumnDefinitions), SIN la celda MEASURE
     * @param measureName medida a calcular (mismo criterio que buildSQLFixed)
     */
    function buildFixedQuery(rowsDefs, colDefs, measureName, state, measuresGrid, atributesGrid) {
        const model = activeModel();
        const source = model && model.source;
        if (!measureName) throw new Error("Añade al menos una medida al informe.");

        const plan = { workspaceId: source.workspaceId, semanticModelId: source.semanticModelId };
        if (rowsDefs.length === 0 || colDefs.length === 0) {
            return { kind: "powerbi-dax", mode: "fixed", text: "(sin cabeceras de filas o columnas)", plan: Object.assign(plan, { empty: true }) };
        }

        const expr = measureExpression(measuresGrid, measureName);
        const filters = (state.Filters || [])
            .filter(f => String(f.Value).trim() !== "")
            .map(f => buildFilterPredicate(atributesGrid, f))
            .filter(Boolean);

        const rowBlocks = axisBlocks(rowsDefs, atributesGrid, "r", "a");
        const colBlocks = axisBlocks(colDefs, atributesGrid, "c", "b");

        let dax = "DEFINE\n";
        filters.forEach((f, i) => { dax += `    VAR __F${i} = FILTER(ALL(${f.columns.join(", ")}), ${f.predicate})\n`; });
        [...rowBlocks, ...colBlocks].forEach(b => { dax += `    VAR ${b.varName} = ${b.table}\n`; });

        const pairs = [];
        rowBlocks.forEach((rb, ri) => colBlocks.forEach((cb, ci) => pairs.push({ ri, ci, rb, cb })));
        const maxKeys = Math.max(...pairs.map(p => new Set([...p.rb.cols, ...p.cb.cols].map(c => c.ref)).size));

        const planPairs = [];
        const parts = pairs.map((p, pi) => {
            // Columnas de agrupación (sin repetir si un campo está en los dos ejes)
            const groupRefs = [];
            [...p.rb.cols, ...p.cb.cols].forEach(c => { if (!groupRefs.includes(c.ref)) groupRefs.push(c.ref); });

            const treatas = b => `TREATAS(SELECTCOLUMNS(${b.varName}, ${b.cols.map(c => `${daxString(c.name)}, [${c.name}]`).join(", ")}), ${b.cols.map(c => c.ref).join(", ")})`;

            const args = [
                ...groupRefs,
                treatas(p.rb),
                treatas(p.cb),
                ...filters.map((f, i) => `__F${i}`),
                `"v", ${expr}`
            ];
            const cols = [
                `"p", ${pi}`,
                ...Array.from({ length: maxKeys }, (_, k) => `"k${k}", ${k < groupRefs.length ? groupRefs[k] : "BLANK()"}`),
                `"v", [v]`
            ];

            planPairs.push({
                groupRefs,
                rowIdx: p.rb.cols.map(c => groupRefs.indexOf(c.ref)),
                colIdx: p.cb.cols.map(c => groupRefs.indexOf(c.ref)),
                rowTypes: p.rb.cols.map(c => c.dtType),
                colTypes: p.cb.cols.map(c => c.dtType),
                rowItems: p.rb.items,
                colItems: p.cb.items
            });

            return "    SELECTCOLUMNS(\n        SUMMARIZECOLUMNS(\n            " + args.join(",\n            ") +
                "\n        ),\n        " + cols.join(",\n        ") + "\n    )";
        });

        dax += "EVALUATE\n" + (parts.length === 1 ? parts[0].replace(/^ {4}/gm, "") : "UNION(\n" + parts.join(",\n") + "\n)");

        return { kind: "powerbi-dax", mode: "fixed", text: dax, plan: Object.assign(plan, { pairs: planPairs }) };
    }

    /**
     * Reparte cada fila agrupada a sus ROW_ID/COLUMN_ID (puede haber
     * cabeceras repetidas: una misma combinación en varias filas/columnas)
     * y devuelve los triples ROW_ID, COLUMN_ID, IMPORTE ordenados, que es
     * lo que lee parseJsonValueTriples.
     */
    function fixedToPseudoBqJson(rows, plan) {
        const lookups = plan.pairs.map(pp => {
            const index = (items) => {
                const m = new Map();
                items.forEach(it => {
                    const k = it.values.join("\u0001");
                    if (!m.has(k)) m.set(k, []);
                    m.get(k).push(it.id);
                });
                return m;
            };
            return { rows: index(pp.rowItems), cols: index(pp.colItems) };
        });

        const cells = new Map();
        rows.forEach(row => {
            const v = cell(row, "v");
            if (isNull(v)) return;
            const pi = Number(cell(row, "p"));
            const pp = plan.pairs[pi];
            if (!pp) return;
            const keyVals = pp.groupRefs.map((_, k) => cell(row, "k" + k));
            const rk = pp.rowIdx.map((ix, i) => normKey(keyVals[ix], pp.rowTypes[i])).join("\u0001");
            const ck = pp.colIdx.map((ix, i) => normKey(keyVals[ix], pp.colTypes[i])).join("\u0001");
            const rIds = lookups[pi].rows.get(rk) || [];
            const cIds = lookups[pi].cols.get(ck) || [];
            rIds.forEach(r => cIds.forEach(c => cells.set(r + "_" + c, { r, c, v })));
        });

        const out = Array.from(cells.values())
            .sort((a, b) => (a.r - b.r) || (a.c - b.c))
            .map(x => '{"f":[{"v": "' + x.r + '"},{"v": "' + x.c + '"},{"v": "' + String(x.v) + '"}]}');
        return '{"rows":[' + out.join(",") + "]}";
    }

    // -----------------------------------------------------------------
    // 2) Ejecución (API executeQueries de Power BI)
    // -----------------------------------------------------------------
    function extractPbiError(body, status) {
        try {
            const e = body && body.error;
            if (e) {
                const details = e["pbi.error"] && e["pbi.error"].details;
                const detail = details && details.find(d => d.detail && d.detail.value);
                if (detail) return detail.detail.value;
                return e.message || e.code || ("HTTP " + status);
            }
            const r = body && body.results && body.results[0];
            if (r && r.error) return r.error.message || JSON.stringify(r.error);
        } catch (err) { /* */ }
        return "HTTP " + status;
    }

    async function postExecuteQueries(plan, dax, token) {
        const path = `/groups/${encodeURIComponent(plan.workspaceId)}/datasets/${encodeURIComponent(plan.semanticModelId)}/executeQueries`;
        const payload = JSON.stringify({
            queries: [{ query: dax }],
            serializerSettings: { includeNulls: true }
        });

        let response;
        try {
            response = await fetch(PBI_API + path, {
                method: "POST",
                headers: { "Authorization": "Bearer " + token, "Content-Type": "application/json" },
                body: payload
            });
        } catch (err) {
            // TypeError = el navegador ha bloqueado la llamada (CORS o red):
            // se repite a través del gateway de Draco, si hay uno configurado.
            if (!(err instanceof TypeError)) throw err;
            const fb = fabricSession();
            if (!fb || !fb.isGatewayConfigured || !fb.isGatewayConfigured()) {
                throw new Error("El navegador ha bloqueado la llamada a la API de Power BI y no hay un gateway configurado (fabricGatewayUrl en config.js).");
            }
            console.warn("[Power BI] Llamada directa bloqueada, se usa el gateway:", err.message);
            response = await fetch(fb.gatewayUrl() + "/powerbi/executeQueries", {
                method: "POST",
                headers: { "Authorization": "Bearer " + token, "Content-Type": "application/json" },
                body: JSON.stringify({ workspaceId: plan.workspaceId, datasetId: plan.semanticModelId, query: dax })
            });
        }

        const text = await response.text();
        let body = null;
        try { body = text ? JSON.parse(text) : null; } catch (e) { body = null; }

        if (response.status === 401) throw new Error("Sin autorización para consultar el modelo. Vuelve a conectar tu conexión de Microsoft Fabric.");
        if (!response.ok || !body || (body.results && body.results[0] && body.results[0].error)) {
            const msg = extractPbiError(body, response.status);
            if (response.status === 403) throw new Error("No tienes permiso para consultar este modelo (hace falta permiso de compilación / Build). " + msg);
            throw new Error("Error en la consulta DAX: " + msg);
        }
        const table = body.results && body.results[0] && body.results[0].tables && body.results[0].tables[0];
        return (table && table.rows) || [];
    }

    // Valor de una fila de executeQueries por alias ("[c0]" o "c0")
    function cell(row, alias) {
        if (Object.prototype.hasOwnProperty.call(row, "[" + alias + "]")) return row["[" + alias + "]"];
        if (Object.prototype.hasOwnProperty.call(row, alias)) return row[alias];
        const k = Object.keys(row).find(key => key.endsWith("[" + alias + "]"));
        return k ? row[k] : null;
    }

    const isNull = v => v === null || v === undefined;

    // Orden ascendente como BigQuery: NULL primero, números como números
    function compareValues(a, b) {
        if (isNull(a) && isNull(b)) return 0;
        if (isNull(a)) return -1;
        if (isNull(b)) return 1;
        if (typeof a === "number" && typeof b === "number") return a - b;
        if (typeof a === "boolean" && typeof b === "boolean") return (a === b) ? 0 : (a ? 1 : -1);
        const sa = String(a), sb = String(b);
        return sa < sb ? -1 : (sa > sb ? 1 : 0);
    }

    /**
     * Claves de orden de un eje: reproduce el ORDER BY del DENSE_RANK de
     * buildFinalSelect / buildAxisOrderByWithSubtotals.
     */
    function axisSortKeys(axisFields, withSubtotalsOnAxis, subtotalsOnTop) {
        const keys = [];
        if (!withSubtotalsOnAxis) {
            axisFields.filter(f => f.hierarchy > 0).forEach(f => keys.push({ type: "value", f }));
            return keys;
        }
        const topVal = subtotalsOnTop ? 0 : 1;
        const bottomVal = subtotalsOnTop ? 1 : 0;
        if (axisFields.length > 0 && axisFields[0].subtotal) {
            keys.push({ type: "firstFlag", f: axisFields[0], topVal, bottomVal });
        }
        axisFields.forEach((f, idx) => {
            if (idx > 0) keys.push({ type: "flag", f, desc: subtotalsOnTop });
            if (f.hierarchy > 0) keys.push({ type: "value", f });
        });
        return keys;
    }

    function keyValue(rec, k) {
        if (k.type === "value") return rec.values[k.f.alias];
        const flag = rec.flags[k.f.alias];
        if (k.type === "firstFlag") return flag === 1 ? k.topVal : k.bottomVal;
        return flag;
    }

    function compareByKeys(a, b, keys) {
        for (const k of keys) {
            let c = compareValues(keyValue(a, k), keyValue(b, k));
            if (k.type === "flag" && k.desc) c = -c;
            if (c !== 0) return c;
        }
        return 0;
    }

    /** DENSE_RANK: asigna 1..n según las claves de orden */
    function denseRank(records, keys, target) {
        const sorted = records.slice().sort((a, b) => compareByKeys(a, b, keys));
        let rank = 0;
        let prev = null;
        sorted.forEach(rec => {
            if (prev === null || compareByKeys(prev, rec, keys) !== 0) rank++;
            rec[target] = rank;
            prev = rec;
        });
    }

    /** Filtro final "NOT (todos los campos de una dimensión IS NULL)" por eje */
    function passesFinalWhere(rec, axisFields, skipAxis) {
        if (skipAxis) return true;
        const byDim = new Map();
        axisFields.filter(f => f.hierarchy > 0).forEach(f => {
            const d = String(f.dimension).toUpperCase();
            if (!byDim.has(d)) byDim.set(d, []);
            byDim.get(d).push(f);
        });
        for (const group of byDim.values()) {
            if (group.every(f => isNull(rec.values[f.alias]))) return false;
        }
        return true;
    }

    function toPseudoBqJson(records, plan) {
        const out = [];
        records.forEach(rec => {
            const vals = [rec.ROW_ID, rec.COLUMN_ID];
            plan.fields.forEach(f => {
                let v = rec.values[f.alias];
                if (plan.withSubtotals && rec.flags[f.alias] === 1) v = "TOTAL";
                vals.push(v);
            });
            plan.measures.forEach(m => vals.push(rec.measures[m.alias]));
            out.push('{"f":[' + vals.map(v => {
                if (isNull(v)) return '{"v": null}';
                const text = String(v).replace(/\\/g, "\\\\").replace(/"/g, '\\"');
                return '{"v": "' + text + '"}';
            }).join(",") + "]}");
        });
        return '{"rows":[' + out.join(",") + "]}";
    }

    /** Post-proceso: mismo resultado que el SELECT final + ORDER BY de buildSQL */
    function shapeResult(rows, plan) {
        const colFields = plan.fields.filter(f => f.axis === "cols");
        const rowFields = plan.fields.filter(f => f.axis === "rows");
        const colsHaveSubtotal = plan.withSubtotals && colFields.some(f => f.subtotal);
        const rowsHaveSubtotal = plan.withSubtotals && rowFields.some(f => f.subtotal);

        let records = rows.map(row => {
            const values = {}, flags = {}, measures = {};
            plan.fields.forEach(f => {
                const present = Number(cell(row, "f" + f.alias.slice(1))) !== 1;
                values[f.alias] = present ? cell(row, f.alias) : null;
                flags[f.alias] = present ? 0 : 1;
            });
            plan.measures.forEach(m => { measures[m.alias] = cell(row, m.alias); });
            return { values, flags, measures };
        });

        records = records.filter(rec =>
            passesFinalWhere(rec, colFields, colsHaveSubtotal) &&
            passesFinalWhere(rec, rowFields, rowsHaveSubtotal));

        denseRank(records, axisSortKeys(colFields, colsHaveSubtotal, plan.subtotalsOnTop), "ROW_ID");
        denseRank(records, axisSortKeys(rowFields, rowsHaveSubtotal, plan.subtotalsOnTop), "COLUMN_ID");

        records.sort((a, b) => (a.ROW_ID - b.ROW_ID) || (a.COLUMN_ID - b.COLUMN_ID));
        return toPseudoBqJson(records, plan);
    }

    /**
     * Ejecuta la consulta devuelta por buildDynamicQuery y devuelve el texto
     * "pseudo JSON de BigQuery" que espera jsonTo3Matrices.
     */
    async function execute(query) {
        const fb = fabricSession();
        if (!fb) {
            throw new Error("[powerBiQuery v5] No se encuentra FB: js/fabric.js no se ha cargado en esta página.");
        }
        if (typeof fb.getPowerBiToken !== "function") {
            throw new Error("[powerBiQuery v5] El js/fabric.js cargado es una versión antigua (sin getPowerBiToken). Probablemente Office tiene la versión anterior en caché.");
        }
        if (query.mode === "fixed" && query.plan.empty) return '{"rows":[]}';
        const token = await fb.getPowerBiToken();
        const rows = await postExecuteQueries(query.plan, query.text, token);
        if (query.mode === "members") return membersToPseudoBqJson(rows, query.plan.names);
        if (query.mode === "fixed") return fixedToPseudoBqJson(rows, query.plan);
        return shapeResult(rows, query.plan);
    }

    const api = {
        version: 5,
        isActiveModel,
        buildDynamicQuery,
        buildMembersQuery,
        buildFixedQuery,
        execute,
        _shapeResult: shapeResult,
        _membersToPseudoBqJson: membersToPseudoBqJson,
        _fixedToPseudoBqJson: fixedToPseudoBqJson
    };
    if (typeof window !== "undefined") window.PowerBIQuery = api;
    if (typeof module !== "undefined") module.exports = api;

})();
