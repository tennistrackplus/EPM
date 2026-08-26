/**
 * ============================================================
 * EPM ADD-IN — CAPA DE ABSTRACCIÓN DE PROVEEDOR
 * ============================================================
 * El resto del add-in habla con "Provider", nunca directamente con
 * BQ o SF, así el explorador del modelo semántico (y, en una fase
 * posterior, el motor de informes) funciona igual sobre BigQuery o
 * sobre Snowflake. Mismo patrón que en Draco Planning.
 */
const Provider = {
    key() {
        return localStorage.getItem("draco_active_provider") || "bigquery";
    },
    setKey(k) {
        localStorage.setItem("draco_active_provider", k);
    },
    label() {
        return this.key() === "snowflake" ? "Snowflake" : "BigQuery";
    },
    /** Etiqueta del primer nivel de la jerarquía de metadatos ("proyecto" en BQ, "base de datos" en SF) */
    level1Label() {
        return this.key() === "snowflake" ? "base de datos" : "proyecto";
    },
    /** Etiqueta del segundo nivel ("dataset" en BQ, "esquema" en SF) */
    level2Label() {
        return this.key() === "snowflake" ? "esquema" : "dataset";
    },

    isConnected() {
        return this.key() === "snowflake" ? SF.isConnected() : BQ.isConnected();
    },

    logout() {
        if (this.key() === "snowflake") {
            SF.logout();
        } else {
            BQ.logout();
        }
    },

    esc(v) {
        return BQ.esc(v);
    },

    toIdentifier(v) {
        return BQ.toIdentifier(v);
    },

    /** Referencia totalmente cualificada a una tabla: proyecto.dataset.tabla (BQ) o BD.esquema.tabla (SF) */
    qualify(c1, c2, table) {
        if (this.key() === "snowflake") {
            return `${c1}.${c2}.${table}`;
        }
        return `\`${c1}.${c2}.${table}\``;
    },

    /** Ejecuta SQL contra el proveedor activo y devuelve {fields:[{name}], rows:[{col: valor}]} */
    async runQuery(sql, level1Id, level2Id) {
        if (this.key() === "snowflake") {
            return SF.runQuerySql(sql, { database: level1Id, schema: level2Id });
        }
        // El "proyecto" usado para lanzar el job de BigQuery es el proyecto
        // de facturación si se indicó uno al crear la conexión (igual que en
        // Power BI); las tablas ya van totalmente cualificadas en el SQL
        // (qualify() incluye siempre el proyecto de los datos), así que esto
        // solo afecta a qué proyecto paga la consulta.
        const billingProject = BQ.getBillingProject() || level1Id;
        return BQ.runQuerySql(billingProject, sql);
    },

    // -----------------------------------------------------------
    // Explorador de metadatos (usado por semantic_model.js)
    // -----------------------------------------------------------
    /** Nivel 1: proyectos (BQ) o bases de datos (SF). Devuelve [{id, label}] */
    async listLevel1() {
        if (this.key() === "snowflake") {
            const names = await SF.listDatabases();
            return names.map(n => ({ id: n, label: n }));
        }
        const projects = await BQ.listProjects();
        return projects.map(p => {
            const id = p.id || (p.projectReference && p.projectReference.projectId);
            return { id, label: id };
        });
    },

    /** Nivel 2: datasets (BQ) o esquemas (SF) dentro del nivel 1 */
    async listLevel2(level1Id) {
        if (this.key() === "snowflake") {
            const names = await SF.listSchemas(level1Id);
            return names.map(n => ({ id: n, label: n }));
        }
        const datasets = await BQ.listDatasets(level1Id);
        return datasets.map(d => {
            const id = d.datasetReference.datasetId;
            return { id, label: id };
        });
    },

    /** Tablas dentro de nivel1/nivel2 */
    async listTables(level1Id, level2Id) {
        if (this.key() === "snowflake") {
            const names = await SF.listTables(level1Id, level2Id);
            return names.map(n => ({ id: n, label: n }));
        }
        const tables = await BQ.listTables(level1Id, level2Id);
        return tables.map(t => {
            const id = t.tableReference.tableId;
            return { id, label: id };
        });
    },

    /** Campos de una tabla: [{name, type}] con el tipo ya normalizado */
    async getTableFields(level1Id, level2Id, tableId) {
        if (this.key() === "snowflake") {
            return SF.getTableFields(level1Id, level2Id, tableId);
        }
        return BQ.getTableFields(level1Id, level2Id, tableId);
    }
};
