/**
 * ============================================================
 * EPM ADD-IN — CAPA DE ABSTRACCIÓN DE PROVEEDOR
 * ============================================================
 * El resto del add-in habla con "Provider", nunca directamente con un
 * motor concreto. Motores: BigQuery (BQ, bigquery.js) y Databricks (DBX,
 * databricks.js, a través de la pasarela del cliente).
 *
 * Jerarquía de metadatos (3 niveles en los dos motores):
 *   BigQuery:   proyecto > dataset > tabla   -> `p.d.t`
 *   Databricks: catálogo > esquema > tabla   -> `c`.`s`.`t`
 *
 * El motor activo es el de la conexión activa (Connections). En los
 * diálogos, que pueden no compartir localStorage con el panel, se usa el
 * que les pasó el panel al compartir la sesión (draco_session_provider).
 */
const Provider = {
    KNOWN: ["bigquery", "databricks"],

    key() {
        try {
            if (typeof Connections !== "undefined" && Connections.getActive) {
                const c = Connections.getActive();
                if (c && this.KNOWN.includes(c.provider)) return c.provider;
            }
        } catch (e) { /* sin conexiones */ }
        try {
            const s = sessionStorage.getItem("draco_session_provider");
            if (this.KNOWN.includes(s)) return s;
        } catch (e) { /* sin acceso */ }
        return "bigquery";
    },
    setKey(k) {
        // La clave real sale de la conexión activa; esto solo la recuerda
        // para los diálogos y limpia el valor heredado de versiones antiguas.
        try { localStorage.removeItem("draco_active_provider"); } catch (e) { /* sin acceso */ }
        try { if (this.KNOWN.includes(k)) sessionStorage.setItem("draco_session_provider", k); } catch (e) { /* sin acceso */ }
    },
    isDatabricks() {
        return this.key() === "databricks";
    },
    /** Módulo de sesión del motor activo (misma interfaz en BQ y DBX). */
    session() {
        return this.isDatabricks() && window.DBX ? window.DBX : window.BQ;
    },
    label() {
        return this.isDatabricks() ? "Databricks" : "BigQuery";
    },
    /** Etiqueta del primer nivel de la jerarquía de metadatos */
    level1Label() {
        return this.isDatabricks() ? "catálogo" : "proyecto";
    },
    /** Etiqueta del segundo nivel */
    level2Label() {
        return this.isDatabricks() ? "esquema" : "dataset";
    },

    isConnected() {
        const s = this.session();
        return !!(s && s.isConnected());
    },

    logout() {
        const s = this.session();
        if (s) s.logout();
    },

    /** ¿Puede escribir (planificación)? En Databricks lo decide Unity Catalog. */
    canWrite() {
        if (this.isDatabricks()) return true;
        return !(window.BQ && typeof BQ.canWriteBigQuery === "function") || BQ.canWriteBigQuery();
    },

    // Los dos motores escapan los literales de texto con barra invertida.
    esc(v) {
        return BQ.esc(v);
    },

    toIdentifier(v) {
        return BQ.toIdentifier(v);
    },

    /** Referencia totalmente cualificada a una tabla */
    qualify(c1, c2, table) {
        if (this.isDatabricks()) {
            const q = (x) => "`" + String(x).replace(/`/g, "``") + "`";
            return q(c1) + "." + q(c2) + "." + q(table);
        }
        return `\`${c1}.${c2}.${table}\``;
    },

    /** Limita un SELECT a n filas */
    limit(selectSql, n) {
        const rows = parseInt(n, 10);
        return `${selectSql} LIMIT ${rows}`;
    },

    /** Tipo de texto para CAST(... AS <tipo>) */
    stringType() {
        return "STRING";
    },

    // -----------------------------------------------------------
    // Dialecto SQL del motor de informes (commands.js)
    // -----------------------------------------------------------
    /** Array literal de enteros: [1, 2] en BigQuery, array(1, 2) en Databricks. */
    arrayLiteral(items) {
        const list = (items || []).join(", ");
        return this.isDatabricks() ? "array(" + list + ")" : "[" + list + "]";
    },

    /**
     * Array con los ids cuya condición se cumple, a partir de pares
     * [{id, cond}]. BigQuery: ARRAY(SELECT .. FROM UNNEST([STRUCT..]) WHERE COND).
     * Databricks: array_compact(array(CASE WHEN cond THEN id END, ...)).
     * Devuelve las líneas SIN sangría inicial (el llamante la añade).
     */
    idArrayLines(pairs, idFieldName, CRLF) {
        if (this.isDatabricks()) {
            return "array_compact(array(" + CRLF
                + pairs.map(p => "    CASE WHEN (" + p.cond + ") THEN " + p.id + " END").join("," + CRLF) + CRLF
                + "))";
        }
        const structs = pairs.map(p => "STRUCT(" + p.id + " AS " + idFieldName + ", (" + p.cond + ") AS COND)");
        return "ARRAY(" + CRLF
            + "    SELECT " + idFieldName + " FROM UNNEST([" + CRLF
            + structs.map(s => "        " + s).join("," + CRLF) + CRLF
            + "    ])" + CRLF
            + "    WHERE COND" + CRLF
            + ")";
    },

    /**
     * Despliega columnas array en filas: [{array, alias}] ->
     * BigQuery:   "<source>, UNNEST(A) AS X, UNNEST(B) AS Y"
     * Databricks: "<source> LATERAL VIEW explode(A) T0 AS X LATERAL VIEW explode(B) T1 AS Y"
     */
    explodeFrom(source, items, sep) {
        const s = sep === undefined ? " " : sep;
        if (this.isDatabricks()) {
            return source + items.map((it, i) => s + "LATERAL VIEW explode(" + it.array + ") DRC_X" + i + " AS " + it.alias).join("");
        }
        return source + "," + items.map(it => s + "UNNEST(" + it.array + ") AS " + it.alias).join(",");
    },

    /** Ejecuta SQL y devuelve {fields:[{name}], rows:[{col: valor}]} */
    async runQuery(sql, level1Id, level2Id, kind) {
        if (window.DracoQueryJournal) {
            return window.DracoQueryJournal.track(sql, () => this._runQuery(sql, level1Id), { kind: kind || "Modelo semántico" });
        }
        return this._runQuery(sql, level1Id);
    },

    async _runQuery(sql, level1Id) {
        if (this.isDatabricks()) return DBX.runQuerySql(sql);
        // El "proyecto" usado para lanzar el job de BigQuery es el proyecto
        // de facturación si se indicó uno al crear la conexión; las tablas
        // ya van totalmente cualificadas en el SQL, así que esto solo
        // afecta a qué proyecto paga la consulta.
        const billingProject = BQ.getBillingProject() || level1Id;
        return BQ.runQuerySql(billingProject, sql);
    },

    // -----------------------------------------------------------
    // Explorador de metadatos (usado por semantic_model.js)
    // -----------------------------------------------------------
    /** Nivel 1: proyectos / catálogos. Devuelve [{id, label}] */
    async listLevel1() {
        if (this.isDatabricks()) {
            return (await DBX.listCatalogs()).map(id => ({ id, label: id }));
        }
        const projects = await BQ.listProjects();
        return projects.map(p => {
            const id = p.id || (p.projectReference && p.projectReference.projectId);
            return { id, label: id };
        });
    },

    /** Nivel 2: datasets / esquemas */
    async listLevel2(level1Id) {
        if (this.isDatabricks()) {
            return (await DBX.listSchemas(level1Id)).map(id => ({ id, label: id }));
        }
        const datasets = await BQ.listDatasets(level1Id);
        return datasets.map(d => {
            const id = d.datasetReference.datasetId;
            return { id, label: id };
        });
    },

    /** Tablas */
    async listTables(level1Id, level2Id) {
        if (this.isDatabricks()) {
            return (await DBX.listTables(level1Id, level2Id)).map(id => ({ id, label: id }));
        }
        const tables = await BQ.listTables(level1Id, level2Id);
        return tables.map(t => {
            const id = t.tableReference.tableId;
            return { id, label: id };
        });
    },

    /** Campos de una tabla: [{name, type}] con el tipo ya normalizado */
    async getTableFields(level1Id, level2Id, tableId) {
        if (this.isDatabricks()) return DBX.getTableFields(level1Id, level2Id, tableId);
        return BQ.getTableFields(level1Id, level2Id, tableId);
    }
};
window.Provider = Provider;
